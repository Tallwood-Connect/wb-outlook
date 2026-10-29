/* Send to Wunderbuild: Outlook task pane. Uses the Wunderbuild public API (beta). */
const APP_VERSION = "1.43"; // bump with every change sent to GitHub; must match data-v in taskpane.html
const DEFAULT_BASE = "https://api.wunderbuild.com/api";
const $ = (id) => document.getElementById(id);
const store = {
  get: (k) => { try { return localStorage.getItem(k); } catch { return null; } },
  set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } },
  remove: (k) => { try { localStorage.removeItem(k); } catch { /* storage unavailable */ } },
  keys: () => { try { return Object.keys(localStorage); } catch { return []; } },
};
// Removes every saved item whose name matches the test.
const purge = (test) => store.keys().filter(test).forEach((k) => store.remove(k));
let job = null; // { _id, name }

function setStatus(msg, kind = "") { const s = $("status"); s.textContent = msg; s.className = kind; }
function base() { return ($("base").value.trim() || DEFAULT_BASE).replace(/\/+$/, ""); }
// The main button: normal, or a green tick when this email is already saved in the job.
let dupFound = false;
function refreshSend() {
  const ready = !!(job && $("name").value.trim());
  const compose = isCompose();
  const attOn = !compose && selectedAttachments().length > 0;
  const btn = $("send");
  if (dupFound && !compose && !(attOn && ready)) { btn.textContent = "Sent to Wunderbuild"; btn.className = "done"; btn.disabled = true; }
  else {
    btn.className = "";
    btn.textContent = dupFound && !compose ? "Save attachments" : compose ? "Save and send" : "Send to Wunderbuild";
    btn.disabled = !ready;
  }
  $("wbAttachBtn").disabled = !job || sending;
  if (!job) $("wbPick").hidden = true;
}

async function api(path, opts = {}) {
  const key = $("key").value.trim();
  if (!key) { $("settings").open = true; throw new Error("Add your Wunderbuild API key under API settings."); }
  let res;
  try {
    res = await fetch(base() + path, {
      ...opts,
      headers: { "x-api-key": key, "Content-Type": "application/json", ...(opts.headers || {}) },
    });
  } catch (err) {
    throw new Error(`Request to ${base()}${path.split("?")[0]} was blocked or failed (${err && err.message}). Check the API base URL and the proxy settings.`);
  }
  if (!res.ok) throw new Error(`Wunderbuild returned ${res.status}: ${(await res.text()).slice(0, 200)}`);
  return res.status === 204 ? null : res.json();
}

/* ---------- conversation record ---------- */
const CONV_PREFIX = "wb_conv_", CONV_MAX_AGE = 90 * 864e5, CONV_MAX_COUNT = 500;

function saveConvRecord(convId, j) {
  if (!convId || !j) return;
  store.set(CONV_PREFIX + convId, JSON.stringify({ at: Date.now(), j: { _id: j._id, name: j.name, jobNumber: j.jobNumber } }));
}
function getConvRecord(convId) {
  if (!convId) return null;
  try {
    const r = JSON.parse(store.get(CONV_PREFIX + convId) || "null");
    if (r && r.j && r.j._id && Date.now() - r.at < CONV_MAX_AGE) return r;
  } catch { /* ignore bad record */ }
  return null;
}
// Drops conversation records older than 90 days, and the oldest ones beyond 500.
function pruneConv() {
  const recs = [];
  for (const k of store.keys().filter((k) => k.startsWith(CONV_PREFIX))) {
    try {
      const r = JSON.parse(store.get(k));
      if (r && r.at && Date.now() - r.at < CONV_MAX_AGE) recs.push({ k, at: r.at }); else store.remove(k);
    } catch { store.remove(k); }
  }
  recs.sort((a, b) => b.at - a.at).slice(CONV_MAX_COUNT).forEach((r) => store.remove(r.k));
}
const fmtDate = (t) => new Date(t).toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" });

/* ---------- the open email (read mode) or draft (compose mode) ---------- */
const isCompose = () => { const it = Office.context.mailbox.item; return !!it && typeof it.subject === "object"; };
let cur = null; // snapshot of the email the panel is showing
const asyncVal = (fn) => new Promise((resolve) => { try { fn((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : undefined)); } catch { resolve(undefined); } });
const newId = () => {
  const a = new Uint8Array(5);
  try { crypto.getRandomValues(a); } catch { for (let i = 0; i < a.length; i++) a[i] = Math.floor(Math.random() * 256); }
  return [...a].map((b) => b.toString(16).padStart(2, "0")).join("");
};
function readSnapshot() { // an email being read: all details are available straight away
  const it = Office.context.mailbox.item;
  if (!it) return null;
  return { compose: false, subject: it.subject || "", from: it.from, to: it.to || [], cc: it.cc || [], bcc: [],
    date: it.dateTimeCreated ? new Date(it.dateTimeCreated) : new Date(), attachments: it.attachments || [],
    conversationId: it.conversationId, internetMessageId: it.internetMessageId, wbId: null, props: null };
}
async function loadDraftSnapshot() { // an email being written: details are read from Outlook one by one
  const it = Office.context.mailbox.item;
  if (!it) return null;
  const [subject, from, to, cc, bcc, atts, props] = await Promise.all([
    asyncVal((cb) => it.subject.getAsync(cb)), asyncVal((cb) => it.from.getAsync(cb)),
    asyncVal((cb) => it.to.getAsync(cb)), asyncVal((cb) => it.cc.getAsync(cb)), asyncVal((cb) => it.bcc.getAsync(cb)),
    asyncVal((cb) => it.getAttachmentsAsync(cb)), asyncVal((cb) => it.loadCustomPropertiesAsync(cb)),
  ]);
  let wbId = props ? props.get("wbId") : null; // a draft's own reference, kept on the draft so saving twice is recognised
  if (!wbId) { wbId = newId(); if (props) { props.set("wbId", wbId); props.saveAsync(() => {}); } }
  const me = Office.context.mailbox.userProfile;
  return { compose: true, subject: subject || "", from: from || (me ? { displayName: me.displayName, emailAddress: me.emailAddress } : null),
    to: to || [], cc: cc || [], bcc: bcc || [], date: new Date(), attachments: atts || [],
    conversationId: undefined, internetMessageId: undefined, wbId, props };
}

// Runs whenever an email is opened: fills in the job from an earlier email in the same
// conversation (saved on this computer), or else from the subject.
let fillRun = 0;
async function smartFill() {
  const run = ++fillRun;
  if (!cur) return;
  const rec = job ? null : getConvRecord(cur.conversationId);
  if (rec) {
    await chooseJob(rec.j);
    if (run !== fillRun) return; // another email was selected meanwhile
    $("autoNote").textContent = `Job filled in from an earlier email in this conversation, saved ${fmtDate(rec.at)}. Check it is correct.`;
  } else {
    await autoMatch();
  }
}

/* ---------- job search ---------- */
let timer;
$("q").addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(searchJobs, 300); });

function renderJobList(jobs) {
  const list = $("jobs");
  list.textContent = "";
  const byName = (a, b) => String(a.name || "").localeCompare(String(b.name || ""), undefined, { numeric: true, sensitivity: "base" });
  for (const j of [...jobs].sort(byName)) {
    const li = document.createElement("li");
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = `${j.jobNumber ? "#" + j.jobNumber + " " : ""}${j.name || "Untitled job"}`;
    if (j.address) { const s = document.createElement("small"); s.textContent = j.address; b.appendChild(s); }
    b.addEventListener("click", () => { jobAuto = false; chooseJob(j); });
    li.appendChild(b); list.appendChild(li);
  }
}

/* ---------- job details in the subject (emails being written only) ---------- */
let appliedPrefix = ""; // the text this add-in last put at the start of the subject
function subjectPrefixOn() { return isCompose() && store.get("wb_subjprefix") === "1"; }
function buildPrefix(j) { // e.g. "J-01067 - Smith (#12 Oak Street, PERTH)"
  const num = j.jobNumber ? "J-" + String(j.jobNumber).replace(/\D/g, "").padStart(5, "0") : "";
  const name = String(j.name || "").replace(/\s*[-–]\s*D\s*$/i, "").trim(); // drops a trailing "- D"
  let street = String(j.address || "").trim();
  if (/^\d/.test(street)) street = "#" + street; // street number gets a #; "Lot 5" and "L5" do not
  const place = [street, String(j.city || "").trim().toUpperCase()].filter(Boolean).join(", ");
  return [num, name].filter(Boolean).join(" - ") + (place ? ` (${place})` : "");
}
function stripPrefix(subject) { // the subject without the job details this add-in added
  const t = subject || "";
  if (appliedPrefix && t.startsWith(appliedPrefix)) return t.slice(appliedPrefix.length).replace(/^[\s:\-–]+/, "");
  // also recognise details added earlier (e.g. before the panel was reopened): "J-01067 - Name (#12 Street, SUBURB) - "
  const m = /^J-\d{5,}(?: - [^()]*?)?(?: \([^()]*\))?(?:\s*-\s*|\s*$)/.exec(t);
  return m ? t.slice(m[0].length) : t;
}
let subjRun = 0, subjBusy = false;
async function applySubjectPrefix(remove) {
  if (!isCompose() || !cur) return;
  const run = ++subjRun;
  $("subjNote").textContent = "";
  try {
    const it = Office.context.mailbox.item;
    const now = await asyncVal((cb) => it.subject.getAsync(cb));
    if (run !== subjRun || now == null) return;
    const rest = stripPrefix(now);
    const prefix = remove || !job ? "" : buildPrefix(job);
    const next = prefix ? `${prefix} - ${rest}` : rest; // ends "- " ready for the person to type
    appliedPrefix = prefix;
    if (next.trimEnd() === now.trimEnd()) return; // Outlook may trim the trailing space
    cur.subject = next; subjBusy = true; // so the live watcher does not treat this as the person typing
    await new Promise((resolve, reject) => it.subject.setAsync(next, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error((r.error && r.error.message) || "Outlook refused")))));
    if (!nameEdited) { $("name").value = suggestName(); refreshSend(); }
  } catch (e) { $("subjNote").textContent = `Could not change the subject: ${e && e.message ? e.message : e}`; }
  finally { if (run === subjRun) subjBusy = false; }
}

let searchRun = 0;
async function searchJobs() {
  const run = ++searchRun, q = $("q").value.trim();
  $("jobs").textContent = "";
  if (!q) { setStatus(""); return; }
  try {
    setStatus("Searching…", "busy");
    const data = await api(`/jobs?page=1&limit=15&search=${encodeURIComponent(q)}`);
    if (run !== searchRun) return; // a newer search is running
    setStatus(data.result.length ? "" : "No jobs found.");
    renderJobList(data.result);
  } catch (e) { if (run === searchRun) setStatus(e.message, "err"); }
}

/* ---------- match the email subject to a job ---------- */
const STOP = new Set(["the","and","for","job","site","residence","house","home","lot","unit","street","st","road","rd","avenue","ave","drive","dr","court","ct","place","pl","lane","ln","way","close","cl","crescent","cres","new","build","project","extension","renovation","stage","re","fw","fwd"]);
const norm = (s) => (s || "").toLowerCase().normalize("NFKD").replace(/[\u0300-\u036f]/g, "").replace(/&/g, " and ").replace(/[^a-z0-9]+/g, " ").trim();
const wordsOf = (s) => norm(s).split(" ").filter(Boolean);
let jobsCache = null, jobsAt = 0, refreshing = null;
const JOBS_KEY = "wb_jobs_v2";
// Keep only what matching needs (job number, address, client). Folders are never read here.
const slim = (j) => ({ _id: j._id, jobNumber: j.jobNumber, name: j.name, address: j.address, city: j.city,
  client: j.client ? { firstName: j.client.firstName, lastName: j.client.lastName, businessName: j.client.businessName } : null });

async function fetchAllJobs() {
  const first = await api("/jobs?page=1&limit=100");
  let all = first.result || [];
  const pages = Math.min((first.pagination && first.pagination.totalPages) || 1, 20);
  if (pages > 1) {
    const rest = await Promise.all(Array.from({ length: pages - 1 }, (_, i) => api(`/jobs?page=${i + 2}&limit=100`)));
    for (const d of rest) all = all.concat(d.result || []);
  }
  return all.filter((j) => !j.isDeleted).map(slim);
}

// Uses the saved copy straight away and refreshes it in the background when older than 10 minutes.
async function loadAllJobs() {
  if (!jobsCache) {
    try {
      const c = JSON.parse(store.get(JOBS_KEY) || "null");
      if (c && Array.isArray(c.jobs) && Date.now() - c.at < 864e5) { jobsCache = c.jobs; jobsAt = c.at; }
    } catch { /* ignore bad saved copy */ }
  }
  if (!jobsCache || Date.now() - jobsAt >= 10 * 60 * 1000) {
    refreshing = refreshing || fetchAllJobs()
      .then((jobs) => { jobsCache = jobs; jobsAt = Date.now(); store.set(JOBS_KEY, JSON.stringify({ at: jobsAt, jobs })); return jobs; })
      .finally(() => { refreshing = null; });
    if (!jobsCache) return refreshing;
    refreshing.catch(() => {});
  }
  return jobsCache;
}

function scoreJob(S, j) {
  let best = 0;
  const has = (phrase) => phrase.length > 0 && (" " + S.norm + " ").includes(" " + phrase + " ");
  const up = (n) => { if (n > best) best = n; };
  // Job number: "Job 1042", "#1042" (strong) or a bare 3+ digit number (medium).
  const n = String(j.jobNumber == null ? "" : j.jobNumber).replace(/\D/g, "").replace(/^0+(?=\d)/, "");
  if (n) {
    if (new RegExp(`(?:job|project|ref|no\\.?|number)\\s*#?\\s*0*${n}(?![0-9])`).test(S.raw)) up(100);
    else if (n.length >= 3 && new RegExp(`#\\s*0*${n}(?![0-9])`).test(S.raw)) up(95);
    else if (n.length >= 3 && new RegExp(`(^|[^0-9])0*${n}(?![0-9])`).test(S.raw)) up(78);
  }
  // Address: written out in full, or house number + street name ("45 Banksia Rd" matches "45 Banksia Road").
  const addr = norm(j.address);
  if (addr.length >= 6 && has(addr)) up(90);
  const aw = wordsOf(j.address);
  const house = aw.find((w) => /^\d+[a-z]?$/.test(w));
  const street = aw.filter((w) => !/^\d/.test(w) && w.length >= 3 && !STOP.has(w));
  if (house && street.length && S.words.has(house) && street.every((w) => S.words.has(w))) up(80);
  // Client name or business name.
  const c = j.client;
  if (c) {
    const full = norm(`${c.firstName || ""} ${c.lastName || ""}`), biz = norm(c.businessName), last = norm(c.lastName);
    if ((full.length >= 5 && has(full)) || (biz.length >= 4 && has(biz))) up(70);
    else if (last.length >= 4 && S.words.has(last)) up(55);
  }
  return best;
}

function matchJobs(subject, jobs) {
  const raw = (subject || "").replace(/^\s*((re|fw|fwd)\s*:\s*)+/i, "").toLowerCase();
  const S = { raw, norm: norm(raw), words: new Set(wordsOf(raw)) };
  return jobs.map((j) => ({ j, s: scoreJob(S, j) })).filter((x) => x.s >= 50).sort((a, b) => b.s - a.s);
}

let matchRun = 0;
let jobAuto = false; // true while the chosen job was filled in from the subject (so a changed subject may replace it)
async function autoMatch() {
  const run = ++matchRun, note = $("matchNote");
  note.textContent = "";
  const item = cur;
  note.className = "";
  if (!$("key").value.trim() || !item || !item.subject) return;
  try {
    note.className = "busy"; note.textContent = "Matching the subject to your jobs…";
    const jobs = await loadAllJobs();
    if (run !== matchRun) return; // a newer email was selected meanwhile
    note.className = ""; note.textContent = "";
    const m = matchJobs(item.subject, jobs);
    if (!m.length) {
      note.textContent = "No job matched this subject. Search for the job above.";
      return;
    }
    const [top, second] = m;
    if (top.s >= 75 && (!second || top.s - second.s >= 15)) {
      if (!job || job._id !== top.j._id) await chooseJob(top.j);
      jobAuto = true;
      $("autoNote").textContent = "Job filled in from the email subject. Check it is correct.";
    } else {
      renderJobList(m.slice(0, 3).map((x) => x.j));
      note.textContent = "Possible jobs from the subject. Pick one:";
    }
  } catch (e) { if (run === matchRun) { note.className = ""; note.textContent = `Could not match jobs: ${e.message}`; } }
}

// Forgets the selected job (and its folders) so one email's job never carries over to the next.
function clearJob() {
  job = null; jobAuto = false; $("subjNote").textContent = "";
  folderRun++; searchRun++;
  $("chosen").textContent = "";
  $("matchNote").className = ""; $("matchNote").textContent = "";
  $("autoNote").textContent = "";
  $("corrNote").hidden = true;
  resetDup();
  $("jobs").textContent = "";
  $("q").value = "";
  for (const id of ["folder", "attFolder"]) {
    const sel = $(id);
    sel.innerHTML = "";
    sel.add(new Option("Choose a job first", ""));
    sel.disabled = true;
  }
  fillAttRowSelects();
  refreshSend();
}

async function chooseJob(j) {
  job = j;
  resetPicker(); // files listed for a previous job must not be attached for this one
  matchRun++; // cancel any matching still running for this email
  $("matchNote").className = ""; $("matchNote").textContent = "";
  $("autoNote").textContent = "";
  $("jobs").textContent = "";
  $("q").value = "";
  $("chosen").textContent = `Job: ${j.jobNumber ? "#" + j.jobNumber + " " : ""}${j.name}`;
  refreshSend();
  await loadFolders();
  if (subjectPrefixOn()) await applySubjectPrefix(); // after the folders, so a problem message is not wiped by them
}

/* ---------- folders ---------- */
let folderRun = 0;
let corrMissing = false;
// Creates the job's Correspondence folder (used when saving) and selects it.
async function createCorrespondence(forJob) {
  await api(`/documents/job/${forJob._id}/folders`, { method: "POST", body: JSON.stringify({ name: "Correspondence" }) });
  const root = await api(`/documents/job/${forJob._id}/folders`);
  const f = (root.subFolders || []).find((x) => /^correspondence$/i.test((x.name || "").trim()));
  if (!f) throw new Error("The Correspondence folder could not be created.");
  const sel = $("folder"), o = new Option(f.name, f._id);
  sel.add(o); sel.value = f._id;
  corrMissing = false; $("corrNote").hidden = true; treeCache = null;
}
async function loadFolders() {
  const run = ++folderRun;
  const sel = $("folder"), af = $("attFolder");
  sel.disabled = true; af.disabled = true; $("send").disabled = true;
  sel.innerHTML = ""; af.innerHTML = "";
  sel.add(new Option("Loading folders…", ""));
  setStatus("Loading folders…", "busy");
  $("corrNote").hidden = true;
  let tree = null, failed = null;
  try { tree = await folderTree(job._id); } catch (e) { failed = e; } // every folder, however deep (Archive folders left out)
  if (run !== folderRun) return; // another job was picked meanwhile
  sel.innerHTML = "";
  sel.add(new Option("Job documents (top level)", ""));
  for (const f of tree || []) if (f.id) sel.add(new Option(f.path, f.id));
  // Emails default to Correspondence (also matched when nested, e.g. "Admin / Correspondence").
  const corr = [...sel.options].find((o) => /(^|\/ )correspondence$/i.test(o.text.trim()));
  if (corr) sel.value = corr.value;
  corrMissing = !corr && !failed; // only claim it is missing when the folder list loaded fine
  $("corrNote").hidden = !corrMissing;
  for (const o of sel.options) af.add(new Option(o.text, o.value));
  // Attachments default to "Working Documents"; otherwise the email's folder.
  const work = [...af.options].find((o) => /(^|\/ )working documents$/i.test(o.text.trim()));
  af.value = work ? work.value : sel.value;
  sel.disabled = false; af.disabled = false;
  fillAttRowSelects();
  if (failed) setStatus(failed.message, "err"); else setStatus("");
  refreshSend();
  if (!failed) checkDuplicate();
}

/* ---------- email -> .eml ---------- */
function getEml() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item || typeof item.getAsFileAsync !== "function")
      return reject(new Error("This version of Outlook cannot export emails to a file (needs Mailbox 1.14)."));
    item.getAsFileAsync((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message)));
  });
}
function b64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function suggestName() {
  const item = cur;
  if (!item) return "";
  const d = item.date;
  const day = d.toISOString().slice(0, 10);
  const subject = (stripPrefix(item.subject) || "No subject").replace(/[\\/:*?"<>|\r\n]+/g, " ").trim().slice(0, 100);
  return `${day} - ${subject}`;
}

/* ---------- email details for the PDF ---------- */
function getBodyText() {
  return new Promise((resolve, reject) =>
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message))));
}
const fmtAddr = (a) => (a ? (a.displayName && a.displayName !== a.emailAddress ? `${a.displayName} <${a.emailAddress}>` : a.emailAddress) : "");
const fmtList = (l) => (Array.isArray(l) ? l.map(fmtAddr).filter(Boolean).join("; ") : "");
function emailMeta(item = cur) {
  const atts = (item.attachments || []).filter((a) => !a.isInline).map((a) => `${a.name} (${Math.max(1, Math.round(a.size / 1024))} KB)`);
  return [
    ["From:", fmtAddr(item.from)],
    ["To:", fmtList(item.to)],
    ["Cc:", fmtList(item.cc)],
    ["Date:", item.date ? item.date.toLocaleString("en-AU") : ""],
    ["Attach:", atts.join("; ")],
  ];
}

/* ---------- formatted PDF (rendered by the Cloudflare Worker) ---------- */
function getBodyHtml() {
  return new Promise((resolve, reject) =>
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message))));
}
function bytesToB64(bytes) {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
const escHtml = (t) => String(t == null ? "" : t).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// Matches each cid: reference in the email to an inline image. Outlook does not give a draft's images their
// content IDs, so images are matched by name first (the id is usually "name.png@code"), then in order of appearance
// (several pasted images are all called "image.png").
function buildCidMap(html, inl) {
  const items = (inl || []).filter((r) => r && r.f), map = new Map(), used = new Set();
  const uri = (r) => `data:${r.f.type};base64,${bytesToB64(r.f.bytes)}`;
  const ids = [...new Set([...String(html || "").matchAll(/cid:([^"'\s)>]+)/gi)].map((m) => m[1].toLowerCase()))];
  const left = [];
  for (const id of ids) {
    const base = id.split("@")[0];
    const r = items.find((x) => !used.has(x) && (x.a.name || "").toLowerCase() === base);
    if (r) { used.add(r); map.set(id, uri(r)); } else left.push(id);
  }
  for (const id of left) { const r = items.find((x) => !used.has(x)); if (r) { used.add(r); map.set(id, uri(r)); } }
  for (const r of items) { const k = (r.a.name || "").toLowerCase(); if (k && !map.has(k)) map.set(k, uri(r)); }
  return map;
}

/* ---------- attach files from Wunderbuild (emails being written only) ---------- */
const MAX_ATTACH = 20 * 1024 * 1024; // keeps clear of Outlook's own attachment limits
let pickRun = 0, pickDocs = [];
const sizeText = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round((n || 0) / 1024))} KB`);
function setPickStatus(msg, cls) { $("pickStatus").textContent = msg || ""; $("pickStatus").className = cls || ""; }
function updatePickGo() { $("pickGo").disabled = sending || ![...$("pickList").querySelectorAll("input")].some((c) => c.checked && !c.disabled); }

// Every folder of the job, however deeply nested, as a flat list of { id, path }. The folder list Wunderbuild
// returns only goes a couple of levels down, so folders without sub-folders listed are asked about one by one.
// Folders with names like Archive, Archived or Superseded are left out, so old versions are not attached by mistake.
const isArchiveName = (n) => /archiv|supersed|supercede|obsolete/i.test(String(n || ""));
let treeCache = null; // { jobId, promise }
function folderTree(jobId) {
  if (treeCache && treeCache.jobId === jobId) return treeCache.promise;
  const p = (async () => {
    const root = await api(`/documents/job/${jobId}/folders`);
    let calls = 0;
    const expand = async (f, depth) => {
      if (isArchiveName(f.name)) return; // no need to look inside
      let subs = f.subFolders || [];
      if (!subs.length && depth < 8 && calls < 300) {
        calls++;
        try { subs = (await api(`/documents/job/${jobId}/folders/${f._id}`)).subFolders || []; } catch { /* leave it as a last-level folder */ }
      }
      f.subFolders = subs;
      await Promise.all(subs.map((x) => expand(x, depth + 1)));
    };
    await Promise.all((root.subFolders || []).map((f) => expand(f, 1)));
    const out = [{ id: "", path: "Job documents (top level)" }];
    const walk = (fs, prefix) => { for (const f of fs || []) { if (isArchiveName(f.name)) continue; out.push({ id: f._id, path: prefix + f.name }); walk(f.subFolders, prefix + f.name + " / "); } };
    walk(root.subFolders, "");
    return out;
  })();
  treeCache = { jobId, promise: p };
  p.catch(() => { if (treeCache && treeCache.promise === p) treeCache = null; });
  return p;
}
// Every file in the job (folders are not included), so file names can be searched. Built once per job.
let indexCache = null, indexProgress = () => {};
function jobFileIndex(jobId) {
  if (indexCache && indexCache.jobId === jobId) return indexCache.promise;
  const p = (async () => {
    const folders = await folderTree(jobId);
    const files = []; let done = 0;
    const queue = [...folders];
    const worker = async () => {
      while (queue.length) {
        const f = queue.shift();
        try { for (const d of await listDocs(jobId, f.id)) files.push({ d, path: f.path }); } catch { files.partial = true; }
        indexProgress(++done, folders.length);
      }
    };
    await Promise.all(Array.from({ length: 5 }, worker));
    return files;
  })();
  indexCache = { jobId, promise: p };
  p.catch(() => { if (indexCache && indexCache.promise === p) indexCache = null; });
  return p;
}
function resetPicker() { $("wbPick").hidden = true; $("pickQ").value = ""; $("pickList").textContent = ""; pickDocs = []; pickRun++; indexCache = null; setPickStatus(""); }

function renderPickRows(items) { // items: [{ d, sub }]
  const list = $("pickList");
  list.textContent = ""; pickDocs = items.map((x) => x.d);
  for (const x of items) {
    const row = document.createElement("div"); row.className = "att";
    const cb = document.createElement("input"); cb.type = "checkbox"; cb.setAttribute("aria-label", `Attach ${x.d.name}`);
    cb.addEventListener("change", updatePickGo);
    const nm = document.createElement("span"); nm.textContent = x.d.name;
    const sz = document.createElement("small"); sz.textContent = x.sub;
    row.appendChild(cb); row.appendChild(nm); row.appendChild(sz); list.appendChild(row);
  }
  updatePickGo();
}
async function loadPickList() {
  const run = ++pickRun, forJob = job;
  $("pickList").textContent = ""; pickDocs = []; updatePickGo();
  if (!forJob) return;
  try {
    setPickStatus("Loading files…", "busy");
    const docs = (await listDocs(forJob._id, $("pickFolder").value)).sort(byName);
    if (run !== pickRun || job !== forJob) return;
    setPickStatus(docs.length ? "" : "No files in this folder.");
    renderPickRows(docs.map((d) => ({ d, sub: sizeText(d.size) })));
  } catch (e) { if (run === pickRun) setPickStatus(e.message, "err"); }
}
const byName = (a, b) => String(a.name).localeCompare(String(b.name), undefined, { numeric: true, sensitivity: "base" });
async function runPickSearch() {
  const q = $("pickQ").value.trim().toLowerCase();
  if (!q) return loadPickList();
  const run = ++pickRun, forJob = job;
  if (!forJob) return;
  $("pickList").textContent = ""; pickDocs = []; updatePickGo();
  try {
    setPickStatus("Reading the job's folders…", "busy");
    indexProgress = (n, total) => { if (run === pickRun) setPickStatus(`Reading the job's folders… ${n} of ${total}`, "busy"); };
    const files = await jobFileIndex(forJob._id);
    if (run !== pickRun || job !== forJob) return;
    const words = q.split(/\s+/);
    const hits = files.filter((x) => words.every((w) => String(x.d.name).toLowerCase().includes(w))).sort((a, b) => byName(a.d, b.d));
    const shown = hits.slice(0, 60);
    setPickStatus(hits.length ? (hits.length > shown.length ? `Showing the first ${shown.length} of ${hits.length} matches. Type more to narrow it down.` : `${hits.length} match${hits.length === 1 ? "" : "es"}.`) : "No files found.");
    if (files.partial) setPickStatus($("pickStatus").textContent + " Some folders could not be read.", hits.length ? "" : "err");
    renderPickRows(shown.map((x) => ({ d: x.d, sub: `${x.path} · ${sizeText(x.d.size)}` })));
  } catch (e) { if (run === pickRun) setPickStatus(e.message, "err"); }
}
async function openPicker() {
  if (!job) return;
  const box = $("wbPick");
  if (!box.hidden) { box.hidden = true; return; }
  const forJob = job, sel = $("pickFolder");
  box.hidden = false; $("pickQ").value = "";
  sel.innerHTML = ""; sel.add(new Option("Loading folders…", "")); sel.disabled = true;
  setPickStatus("Loading folders…", "busy"); $("pickList").textContent = ""; pickDocs = []; updatePickGo();
  try {
    const folders = await folderTree(forJob._id);
    if (job !== forJob) return;
    sel.innerHTML = "";
    for (const f of folders) sel.add(new Option(f.path, f.id));
    const work = folders.find((f) => /(^|\/ )working documents$/i.test(f.path.trim()));
    sel.value = work ? work.id : "";
    sel.disabled = false;
    await loadPickList();
  } catch (e) { setPickStatus(e.message, "err"); sel.innerHTML = ""; sel.disabled = false; }
}
// Downloads a file from a signed link, through the proxy when the storage site refuses the browser (CORS).
async function fetchBytes(url) {
  let res;
  try { res = await fetch(url); }
  catch { res = await fetch(base().replace(/\/api$/, "") + "/download", { headers: { "x-api-key": $("key").value.trim(), "x-download-url": url } }); }
  if (!res.ok) throw new Error(`download failed (${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}
function addAttachment(b64, name) {
  return new Promise((resolve, reject) => {
    const it = Office.context.mailbox.item;
    if (!it || typeof it.addFileAttachmentFromBase64Async !== "function") return reject(new Error("this version of Outlook cannot attach files from the add-in"));
    it.addFileAttachmentFromBase64Async(b64, name, { isInline: false }, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error((r.error && r.error.message) || "Outlook refused"))));
  });
}
async function attachPicked() {
  const forJob = job;
  const rows = [...$("pickList").children];
  const chosen = rows.map((row, i) => ({ row, cb: row.children[0], d: pickDocs[i] })).filter((x) => x.cb.checked && !x.cb.disabled);
  if (!forJob || !chosen.length || sending) return;
  sending = true; lockUI(true); $("pickFolder").disabled = true;
  const failed = []; let done = 0;
  try {
    for (const x of chosen) {
      setPickStatus(`Attaching ${done + failed.length + 1} of ${chosen.length}: ${x.d.name}…`, "busy");
      try {
        if ((x.d.size || 0) > MAX_ATTACH) throw new Error(`too large (${sizeText(x.d.size)}; the limit is ${sizeText(MAX_ATTACH)})`);
        const info = await api(`/documents/job/${forJob._id}/${x.d._id}/download-url`);
        const bytes = await fetchBytes(info.url);
        if (bytes.length > MAX_ATTACH) throw new Error(`too large (${sizeText(bytes.length)})`);
        await addAttachment(bytesToB64(bytes), info.name || x.d.name);
        done++; x.cb.checked = false; x.cb.disabled = true; x.row.children[1].textContent = `${x.d.name} (attached)`;
      } catch (e) { failed.push(`${x.d.name}: ${e.message}`); }
    }
  } finally {
    sending = false; lockUI(false); $("pickFolder").disabled = false; refreshSend(); updatePickGo();
    if (pendingRefresh) { pendingRefresh = false; refreshForItem(true); }
  }
  setPickStatus(`${done ? `Attached ${done} file${done === 1 ? "" : "s"} to your email. ` : ""}${failed.length ? `Not attached: ${failed.join("; ")}` : ""}`, failed.length ? "err" : "ok");
}

// Email header block + the email's own HTML. Scripts are removed, and inline (cid:) images are embedded.
function buildEmailHtml({ subject, meta, bodyHtml, cid }) {
  let body = String(bodyHtml || "").replace(/<script[\s\S]*?<\/script>/gi, "");
  body = body.replace(/cid:([^"'\s)>]+)/gi, (m, id) => (cid && (cid.get(id.toLowerCase()) || cid.get(id.split("@")[0].toLowerCase()))) || m);
  const rows = meta.filter(([, v]) => v).map(([k, v]) =>
    `<tr><td style="padding:1px 10px 1px 0;font-weight:bold;vertical-align:top;white-space:nowrap">${escHtml(k)}</td><td style="padding:1px 0">${escHtml(v)}</td></tr>`).join("");
  return `<!doctype html><html><head><meta charset="utf-8">` +
    `<meta http-equiv="Content-Security-Policy" content="script-src 'none'; object-src 'none'; frame-src 'none'">` +
    `<style>img{max-width:100%!important;height:auto!important}</style></head><body style="margin:0">` + // keeps pictures in proportion
    `<div style="font:13px Arial,Helvetica,sans-serif;color:#000;border-bottom:1px solid #999;padding-bottom:10px;margin-bottom:14px">` +
    `<div style="font-size:18px;font-weight:bold;margin-bottom:8px">${escHtml(subject || "(No subject)")}</div><table style="border-collapse:collapse">${rows}</table></div>` +
    `<div>${body}</div></body></html>`;
}

async function renderPdf(html) {
  const b = base();
  if (/\/\/api\.wunderbuild\.com/i.test(b)) throw new Error("no Worker proxy in the API base URL");
  const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), 60000);
  try {
    let res;
    try {
      res = await fetch(b.replace(/\/api$/, "") + "/pdf", {
        method: "POST", signal: ctl.signal,
        headers: { "x-api-key": $("key").value.trim(), "Content-Type": "application/json" },
        body: JSON.stringify({ html }),
      });
    } catch (e) { throw new Error(e.name === "AbortError" ? "timed out" : "could not reach the Worker"); }
    if (!res.ok) throw new Error((await res.text()).slice(0, 160) || `status ${res.status}`);
    const bytes = new Uint8Array(await res.arrayBuffer());
    if (String.fromCharCode(...bytes.slice(0, 4)) !== "%PDF") throw new Error("the reply was not a PDF");
    return bytes;
  } finally { clearTimeout(timer); }
}

/* ---------- unique reference ID and duplicate check ---------- */
// A short ID derived from the email's own message ID, so the same email gets the same ID for everyone.
// Small non-cryptographic hash (cyrb53 style): gives the same result in every browser, with no extra libraries.
function hash40(str) {
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return ((h2 >>> 0).toString(16).padStart(8, "0") + (h1 >>> 0).toString(16).padStart(8, "0")).slice(0, 10);
}
async function emailId(item) {
  // Runs straight away (no waiting), so it reads the email that is open right now.
  return hash40((item.internetMessageId
    ? String(item.internetMessageId).trim()
    : item.wbId ? String(item.wbId)
    : [fmtAddr(item.from), item.date ? item.date.toISOString() : "", item.subject || ""].join("|")
  ).toLowerCase());
}

async function listDocs(jobId, folderId) {
  const out = [];
  for (let page = 1; page <= 5; page++) {
    const d = await api(`/documents/job/${jobId}?page=${page}&limit=100${folderId ? `&folderId=${folderId}` : ""}`);
    out.push(...(d.result || []).filter((x) => !x.isDeleted));
    if (!d.pagination || !d.pagination.hasNextPage) break;
  }
  return out;
}

// Looks in the folder the email will be saved to, and in Correspondence, plus each one's "Outlook Files" folder.
async function findExisting(forJob, destFolderId, id) {
  const corr = [...$("folder").options].find((o) => /(^|\/ )correspondence$/i.test(o.text.trim()));
  const folders = [...new Set([destFolderId, corr ? corr.value : destFolderId])];
  const lists = await Promise.all(folders.map(async (f) => {
    const [docs, outlookId] = await Promise.all([listDocs(forJob._id, f), findOutlookFolderId(f)]);
    return [...docs, ...(outlookId ? await listDocs(forJob._id, outlookId) : [])];
  }));
  return lists.flat().filter((d) => (d.name || "").includes(`[${id}]`));
}

let dupRun = 0, justSaved = null; // justSaved = id of the job this email was just saved to
function resetDup() { dupRun++; dupFound = false; $("dupNote").className = ""; $("dupNote").textContent = ""; }
async function checkDuplicate() {
  resetDup();
  const run = dupRun, forJob = job, item = cur, note = $("dupNote");
  if (!forJob || !item || justSaved === forJob._id) return; // no "already saved" note straight after saving it
  note.className = "busy"; note.textContent = "Checking whether this email is already saved…";
  try {
    const found = await findExisting(forJob, $("folder").value, await emailId(item));
    if (run !== dupRun || job !== forJob) return; // something changed meanwhile
    note.className = ""; note.textContent = "";
    if (found.length) {
      dupFound = true; // received emails show a green "Sent to Wunderbuild" button instead of a message
      for (const row of $("attList").children) row.children[0].checked = false; // already filed: tick one again to save it separately
      if (isCompose()) note.textContent = "This email was already saved to Wunderbuild. Save and send will only send it.";
      refreshSend();
    }
  } catch (e) { if (run === dupRun) { note.className = ""; note.textContent = `Could not check for an existing copy: ${e.message}`; } }
}

/* ---------- attachments ---------- */
let currentAtts = [];
function renderAttachments() {
  // An email being written already carries its attachments (they are kept in the saved .eml), so only received emails offer to save them.
  currentAtts = (cur && !cur.compose ? cur.attachments || [] : []).filter((a) => !a.isInline);
  const box = $("attSection"), list = $("attList");
  list.innerHTML = "";
  box.hidden = currentAtts.length === 0;
  currentAtts.forEach((a) => {
    const kb = Math.max(1, Math.round((a.size || 0) / 1024));
    const row = document.createElement("div");
    row.className = "att";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.setAttribute("aria-label", `Save ${a.name}`);
    // Tick everything except tiny images (usually signature logos).
    cb.checked = !(attType(a).startsWith("image/") && (a.size || 0) < 20480);
    const nm = document.createElement("input");
    // The file extension is shown (in green) at the right of the box and kept as it is, so only the name is edited.
    const full = a.name || "attachment";
    const em = /\.[A-Za-z0-9]{1,10}$/.exec(full);
    const extText = em && em.index > 0 ? em[0] : "";
    nm.type = "text"; nm.value = extText ? full.slice(0, -extText.length) : full; nm.spellcheck = false;
    if (extText) nm.style.paddingRight = `${extText.length + 2}ch`;
    nm.setAttribute("aria-label", `File name for ${a.name}`);
    const sz = document.createElement("small");
    sz.textContent = `${kb} KB`;
    cb.addEventListener("change", refreshSend);
    const fs = document.createElement("select");
    fs.setAttribute("aria-label", `Folder for ${a.name}`);
    fs.hidden = true; // shown when each attachment has its own folder
    fs.addEventListener("change", () => { fs.dataset.touched = "1"; });
    row.appendChild(cb); row.appendChild(nm); row.appendChild(sz); row.appendChild(fs);
    if (extText) { const ex = document.createElement("span"); ex.className = "ext"; ex.textContent = extText; row.appendChild(ex); }
    list.appendChild(row);
  });
  $("attMode").checked = false; // one folder for all is the standard
  syncAttMode();
  fillAttRowSelects();
}
// Every attachment's folder list is a copy of the "Save attachments to" list.
function fillAttRowSelects() {
  const src = $("attFolder");
  for (const row of $("attList").children) {
    const sel = row.children[3];
    if (!sel) continue;
    const keep = sel.dataset.touched ? sel.value : null;
    sel.innerHTML = "";
    for (const o of src.options) sel.add(new Option(o.text, o.value));
    sel.value = keep !== null && [...sel.options].some((o) => o.value === keep) ? keep : src.value;
    sel.disabled = src.disabled;
  }
}
function syncAttMode() {
  const each = $("attMode").checked;
  $("swOne").classList.toggle("on", !each); $("swEach").classList.toggle("on", each);
  $("attFolder").hidden = each; $("attFolderLabel").hidden = each;
  for (const row of $("attList").children) if (row.children[3]) row.children[3].hidden = !each;
}
// Rows are in the same order as currentAtts: [checkbox, name box, size].
function selectedAttachments() {
  return [...$("attList").children]
    .map((row, i) => ({ att: currentAtts[i], on: row.children[0].checked, typed: row.children[1].value, cb: row.children[0], folder: row.children[3] ? row.children[3].value : "" }))
    .filter((r) => r.on);
}
// Cleans the typed name and keeps the original file type if the new name has no extension.
function finalName(typed, original) { // the typed name plus the original file's extension
  const ext = (original.match(/\.[A-Za-z0-9]{1,10}$/) || [""])[0];
  const stem = ext ? original.slice(0, -ext.length) : original;
  const n = (typed || "").trim().replace(/[\\/:*?"<>|]+/g, "_") || stem;
  return ext && n.toLowerCase().endsWith(ext.toLowerCase()) ? n : n + ext;
}
// Emails being written do not report an attachment's type, so it is worked out from the file extension.
const EXT_TYPES = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", jpe: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff", pdf: "application/pdf" };
function attType(a) {
  if (a.contentType) return a.contentType;
  const m = /\.([a-z0-9]+)$/i.exec(a.name || "");
  return (m && EXT_TYPES[m[1].toLowerCase()]) || "application/octet-stream";
}
function getAttachment(a) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(a.id, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(r.error.message));
      const v = r.value, name = (a.name || "attachment").replace(/[\\/:*?"<>|]+/g, "_");
      if (v.format === "base64") return resolve({ name, type: attType(a), bytes: b64ToBytes(v.content) });
      if (v.format === "eml") return resolve({ name: /\.eml$/i.test(name) ? name : name + ".eml", type: "message/rfc822", bytes: new TextEncoder().encode(v.content) });
      if (v.format === "iCalendar") return resolve({ name: /\.ics$/i.test(name) ? name : name + ".ics", type: "text/calendar", bytes: new TextEncoder().encode(v.content) });
      reject(new Error("it is a cloud link rather than a file"));
    });
  });
}

/* ---------- upload ---------- */
async function uploadFile(name, type, bytes, folderId) {
  let docId = null;
  try {
    const doc = await api(`/documents/job/${job._id}`, {
      method: "POST",
      body: JSON.stringify({ name, type, size: bytes.length, ...(folderId ? { folderId } : {}) }),
    });
    docId = doc._id;
    if (!doc.uploadUrl) throw new Error("Wunderbuild did not return an upload URL.");
    const blob = new Blob([bytes], { type });
    let put;
    try {
      put = await fetch(doc.uploadUrl, { method: "PUT", headers: { "Content-Type": type }, body: blob });
    } catch {
      // Storage blocked the browser (CORS): send the file through the proxy instead.
      put = await fetch(base().replace(/\/api$/, "") + "/upload", {
        method: "PUT",
        headers: { "x-api-key": $("key").value.trim(), "x-upload-url": doc.uploadUrl, "Content-Type": type },
        body: blob,
      });
    }
    if (!put.ok) throw new Error(`Upload of ${name} failed (${put.status}).`);
    return doc;
  } catch (e) {
    if (docId) { try { await api(`/documents/job/${job._id}/${docId}`, { method: "DELETE" }); } catch { /* best effort */ } }
    throw e;
  }
}

// Finds (or creates) a sub-folder called "Outlook Files" inside parentId (or the job root when empty).
async function findOutlookFolderId(parentId) {
  const detail = await api(parentId ? `/documents/job/${job._id}/folders/${parentId}` : `/documents/job/${job._id}/folders`);
  const hit = (detail.subFolders || []).find((f) => (f.name || "").trim().toLowerCase() === "outlook files");
  return hit ? hit._id : null;
}
async function ensureOutlookFolder(parentId) {
  const existing = await findOutlookFolderId(parentId);
  if (existing) return existing;
  const created = await api(`/documents/job/${job._id}/folders`, {
    method: "POST",
    body: JSON.stringify({ name: "Outlook Files", ...(parentId ? { parentId } : {}) }),
  });
  treeCache = null; // the folder list has changed
  return created._id;
}

$("clearCache").addEventListener("click", () => {
  jobsCache = null; jobsAt = 0;
  purge((k) => k === JOBS_KEY || k === "wb_job" || k.startsWith("wb_folder_") || k.startsWith(CONV_PREFIX));
  $("clearNote").textContent = "Saved job list and conversation history cleared.";
});
$("signOut").addEventListener("click", () => {
  jobsCache = null; jobsAt = 0;
  purge((k) => k.startsWith("wb_"));
  $("key").value = ""; $("base").value = DEFAULT_BASE; $("fmtPdf").checked = true; $("askSend").checked = false; $("subjPrefix").checked = false; syncSwitch();
  try { const rs = Office.context.roamingSettings; rs.set("wb_ask", false); rs.saveAsync(() => {}); } catch { /* ignore */ }
  clearJob();
  $("clearNote").textContent = "Your key and all saved data were removed from this computer.";
});

let sending = false, pendingRefresh = false;
function lockUI(on) {
  $("q").disabled = on; $("name").disabled = on; $("wbAttachBtn").disabled = on || !job; $("wbClearBtn").disabled = on; $("pickGo").disabled = on;
  $("folder").disabled = on || !job; $("attFolder").disabled = on || !job;
  for (const el of $("attList").querySelectorAll("input,select")) el.disabled = on;
  for (const el of $("jobs").querySelectorAll("button")) el.disabled = on;
}

// Rebuilds an .eml from a draft (Outlook cannot export a draft). Cloud-linked attachments cannot be included.
async function buildDraftEml(snap, bodyHtml, bodyText, readAtt) {
  const parts = []; let skipped = 0;
  for (const a of snap.attachments) {
    const r = await readAtt(a);
    if (r.f) parts.push({ name: r.f.name || a.name, type: r.f.type, bytes: r.f.bytes, isInline: a.isInline });
    else skipped++;
  }
  return { bytes: buildEml({ id: snap.wbId, from: snap.from, to: snap.to, cc: snap.cc, bcc: snap.bcc, subject: snap.subject,
    date: new Date(), bodyHtml, bodyText, attachments: parts }), skipped };
}
// Sends the draft that was just saved. Only available in newer Outlook (Mailbox 1.15 and later).
function sendDraft() {
  return new Promise((resolve, reject) => {
    const it = Office.context.mailbox.item;
    if (!it || typeof it.sendAsync !== "function") return reject(new Error("unsupported"));
    try { it.sendAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error((r.error && r.error.message) || "Outlook could not send")))); }
    catch (e) { reject(e); }
  });
}
function markSaved(snap) { // lets the send-time prompt know this draft has been saved
  return new Promise((resolve) => {
    if (!snap.compose || !snap.props) return resolve();
    try { snap.props.set("wbSaved", true); snap.props.saveAsync(() => resolve()); } catch { resolve(); }
  });
}

$("folder").addEventListener("change", () => { if (job && !sending) checkDuplicate(); });

$("send").addEventListener("click", async () => {
  if (sending) return;
  sending = true; lockUI(true); $("send").disabled = true;
  try {
    const baseName = $("name").value.trim().replace(/\.(eml|pdf|msg)$/i, "");
    if (corrMissing && job) { setStatus("Creating Correspondence folder…", "busy"); await createCorrespondence(job); }
    const folderId = $("folder").value, folderLabel = $("folder").selectedOptions[0].text;
    const attFolderId = $("attFolder").value, attFolderLabel = $("attFolder").selectedOptions[0].text;
    const atts = selectedAttachments();
    const problems = [];

    // Read everything from the open email first, in one go, so clicking another email cannot mix things up.
    setStatus(isCompose() ? "Reading your draft…" : "Reading email…", "busy");
    const snap = isCompose() ? await loadDraftSnapshot() : readSnapshot(); // a read email is captured instantly
    const subject = snap.subject, meta = emailMeta(snap);
    const convId = snap.conversationId, jobAtSend = job;
    const idPromise = emailId(snap); // identity of the email that is open right now
    const useFmt = $("fmtPdf").checked;
    const memo = new Map();
    const readAtt = (a) => { if (!memo.has(a.id)) memo.set(a.id, getAttachment(a).then((f) => ({ f }), (e) => ({ e }))); return memo.get(a.id); };
    const perFile = $("attMode").checked; // a folder chosen for each attachment, or one folder for all
    const reads = atts.map(({ att, typed, cb, folder }) => readAtt(att).then((r) => ({ att, typed, cb, folder, f: r.f, e: r.e })));
    const inlineAtts = snap.attachments.filter((a) => a.isInline && attType(a).startsWith("image/") && (a.size || 0) < 5 * 1024 * 1024);
    const inlineReads = useFmt ? inlineAtts.map((a) => readAtt(a).then((r) => (r.f ? { a, f: r.f } : null))) : [];
    const [bodyText, bodyHtml, emlB64, ...rest] = await Promise.all([getBodyText(), (useFmt || snap.compose) ? getBodyHtml() : Promise.resolve(""), snap.compose ? null : getEml(), ...reads, ...inlineReads]);
    const got = rest.slice(0, reads.length), inl = rest.slice(reads.length);
    let eml;
    if (snap.compose) {
      const built = await buildDraftEml(snap, bodyHtml, bodyText, readAtt);
      eml = built.bytes;
      if (built.skipped) problems.push(`${built.skipped} attachment${built.skipped === 1 ? "" : "s"} could not be included in the rebuilt email file (cloud links)`);
    } else eml = b64ToBytes(emlB64);

    // Check, right now, whether this email is already in the job (e.g. a colleague saved it first).
    const id = await idPromise;
    let existing = [];
    try { setStatus("Checking for an existing copy…", "busy"); existing = await findExisting(jobAtSend, folderId, id); }
    catch (e) { problems.push(`could not check for an existing copy (${e.message})`); }
    const duplicate = existing.length > 0;
    const fileBase = `${baseName} [${id}]`;

    let pdf = null;
    if (useFmt && !duplicate) {
      try {
        setStatus("Rendering formatted PDF…", "busy");
        const cid = buildCidMap(bodyHtml, inl);
        pdf = await renderPdf(buildEmailHtml({ subject, meta, bodyHtml, cid }));
      } catch (e) { problems.push(`formatted PDF unavailable (${e.message}), so a basic text PDF was saved instead`); }
    }
    if (!duplicate) {
      if (!pdf) pdf = emailToPdf({ subject, meta, body: bodyText });
      setStatus("Saving PDF…", "busy");
      await uploadFile(`${fileBase}.pdf`, "application/pdf", pdf, folderId);

      try {
        setStatus("Saving original email to Outlook Files…", "busy");
        const outlookFolderId = await ensureOutlookFolder(folderId);
        await uploadFile(`${fileBase}.eml`, "message/rfc822", eml, outlookFolderId);
      } catch (e) { problems.push(`original email not saved (${e.message})`); }
    }

    let savedAtts = 0, n = 0;
    for (const r of got) {
      if (!r.f) { problems.push(`${r.att.name}: ${r.e.message}`); continue; }
      try {
        setStatus(`Saving attachment ${++n} of ${atts.length}…`, "busy");
        await uploadFile(finalName(r.typed, r.f.name), r.f.type, r.f.bytes, perFile ? r.folder : attFolderId);
        savedAtts++; r.cb.checked = false; // done: no need to save it again
      } catch (e) { problems.push(`${r.att.name}: ${e.message}`); }
    }

    // Remember which job this conversation was filed to (kept on this computer only).
    saveConvRecord(convId, jobAtSend);
    await markSaved(snap);

    let msg;
    if (duplicate) {
      msg = `This email is already saved in this job (${existing[0].name}), so the PDF and original email were not saved again.`;
      if (!atts.length) msg += " Nothing new to save.";
    } else {
      justSaved = jobAtSend._id;
      msg = `Saved. PDF in ${folderLabel}.`;
      if (!problems.some((p) => p.startsWith("original"))) msg += ` Original email in ${folderLabel} / Outlook Files.`;
    }
    if (atts.length) msg += ` ${savedAtts} of ${atts.length} attachment${atts.length === 1 ? "" : "s"} in ${perFile ? "the chosen folders" : attFolderLabel}.`;
    if (problems.length) msg += ` Problems: ${problems.join("; ")}.`;
    if (!snap.compose && !problems.length) dupFound = true; // everything saved: the button turns into the green "Sent to Wunderbuild"
    const wantSend = snap.compose; // a draft is always sent as it is saved, so the saved copy matches what went out
    if (wantSend && problems.length) {
      msg += " The email has NOT been sent because of the problems above. Fix them, or send it yourself.";
    } else if (wantSend) {
      setStatus(msg + " Sending…", "busy");
      try { await sendDraft(); msg += " Email sent."; }
      catch (e) {
        msg += e.message === "unsupported" ? " This version of Outlook cannot send from the add-in, so click Send yourself."
          : ` The email could not be sent automatically (${e.message}). Click Send yourself.`;
        problems.push("not sent");
      }
    }
    setStatus(msg, problems.length ? "err" : "ok");
  } catch (e) {
    setStatus(e.message, "err");
  } finally {
    sending = false; lockUI(false); refreshSend();
    if (pendingRefresh) { pendingRefresh = false; refreshForItem(true); }
  }
});

// Called when a different email is selected in the pinned panel (or the draft is refreshed).
// Reset also empties the email's attachments (pictures pasted into the body are not attachments and stay).
async function removeAllAttachments() {
  let n = 0;
  try {
    const it = Office.context.mailbox.item;
    const atts = await asyncVal((cb) => it.getAttachmentsAsync(cb));
    for (const a of atts || []) if (!a.isInline) { await new Promise((resolve) => it.removeAttachmentAsync(a.id, (r) => { if (r && r.status === Office.AsyncResultStatus.Succeeded) n++; resolve(); })); }
  } catch { /* nothing to remove, or Outlook refused */ }
  return n;
}

let snapRun = 0;
let nameEdited = false; // true once the person types in the file name box themselves

// Outlook has no "subject changed" event, so while a draft's options are showing the subject is
// checked every couple of seconds and the file name follows it (until the name is edited by hand).
let subjectTimer = null;
function startSubjectWatch() {
  if (subjectTimer) return;
  subjectTimer = setInterval(async () => {
    if (sending || subjBusy || !cur || !isCompose()) return;
    try {
      const it = Office.context.mailbox.item;
      const subject = await asyncVal((cb) => it.subject.getAsync(cb));
      if (subject == null || sending || subject.trimEnd() === cur.subject.trimEnd()) return;
      cur.subject = subject;
      if (!nameEdited) { $("name").value = suggestName(); refreshSend(); }
      // the subject changed: suggest jobs again, unless a job was picked by hand
      if (!job || jobAuto) { if (job) clearJob(); autoMatch(); }
    } catch { /* draft closed or Outlook busy: try again next time */ }
  }, 2000);
}
async function refreshForItem(keepStatus) {
  const run = ++snapRun;
  justSaved = null;
  clearJob();
  const snap = await loadSnapshot();
  if (run !== snapRun) return; // another email was selected meanwhile
  cur = snap;
  nameEdited = false; appliedPrefix = "";
  $("name").value = suggestName();
  renderAttachments();
  if (!keepStatus) setStatus("");
  refreshSend();
  smartFill();
}
const loadSnapshot = () => (isCompose() ? loadDraftSnapshot() : Promise.resolve(readSnapshot()));

/* ---------- init ---------- */
$("name").addEventListener("input", () => { nameEdited = true; refreshSend(); });
["key", "base"].forEach((id) => $(id).addEventListener("change", () => store.set("wb_" + id, $(id).value.trim())));
$("askSend").addEventListener("change", () => {
  try { const rs = Office.context.roamingSettings; rs.set("wb_ask", $("askSend").checked); rs.saveAsync(() => {}); } catch { /* settings unavailable */ }
});
$("refreshDraft").addEventListener("click", async () => {
  if (sending) return;
  await applySubjectPrefix(true); // take the added job details back out of the subject first
  await removeAllAttachments();
  resetPicker();
  refreshForItem(false);
});
function syncSwitch() { $("swAuto").classList.toggle("on", $("subjPrefix").checked); $("swManual").classList.toggle("on", !$("subjPrefix").checked); }
$("subjPrefix").addEventListener("change", () => {
  syncSwitch();
  store.set("wb_subjprefix", $("subjPrefix").checked ? "1" : "0");
  applySubjectPrefix(!$("subjPrefix").checked); // switching it off takes the added text back out
});
$("attMode").addEventListener("change", () => {
  syncAttMode();
  if ($("attMode").checked) for (const row of $("attList").children) { // start each one at the folder already chosen above
    const sel = row.children[3];
    if (sel && !sel.dataset.touched) sel.value = $("attFolder").value;
  }
});
$("wbClearBtn").addEventListener("click", async () => {
  if (sending) return;
  sending = true; lockUI(true);
  try {
    const n = await removeAllAttachments();
    setStatus(n ? `Removed ${n} attachment${n === 1 ? "" : "s"} from the email.` : "There were no attachments to remove.", n ? "ok" : "");
    if (!$("wbPick").hidden) { $("pickQ").value.trim() ? runPickSearch() : loadPickList(); } // lets the files be attached again
  } finally {
    sending = false; lockUI(false); refreshSend();
    if (pendingRefresh) { pendingRefresh = false; refreshForItem(true); }
  }
});
$("wbAttachBtn").addEventListener("click", () => { if (!sending) openPicker(); });
$("pickFolder").addEventListener("change", () => { if (!sending) { $("pickQ").value = ""; loadPickList(); } });
let pickTimer = null;
$("pickQ").addEventListener("input", () => { clearTimeout(pickTimer); pickTimer = setTimeout(() => { if (!sending) runPickSearch(); }, 300); });
$("pickGo").addEventListener("click", attachPicked);
$("fmtPdf").addEventListener("change", () => store.set("wb_fmtpdf", $("fmtPdf").checked ? "1" : "0"));

// Shows the controls that belong to the item in front of the person: a received email or one being written.
function applyMode() {
  const compose = isCompose();
  $("subjBar").hidden = !compose; syncSwitch();
  $("refreshDraft").hidden = !compose;
  $("wbAttach").hidden = !compose;
  refreshSend();
  if (compose) startSubjectWatch();
}

Office.onReady(async () => {
  const ver = $("ver");
  ver.textContent = `v${APP_VERSION}` + (ver.dataset.v === APP_VERSION ? "" : ` (taskpane.html is v${ver.dataset.v}, files out of sync)`);
  purge((k) => k === "wb_job" || k === "wb_jobs_v1" || k.startsWith("wb_folder_")); // leftovers from older versions
  pruneConv();
  $("key").value = store.get("wb_key") || "";
  $("base").value = store.get("wb_base") || DEFAULT_BASE;
  $("fmtPdf").checked = store.get("wb_fmtpdf") !== "0";
  $("subjPrefix").checked = store.get("wb_subjprefix") === "1";
  try { $("askSend").checked = Office.context.roamingSettings.get("wb_ask") === true; } catch { /* settings unavailable */ }
  if (!$("key").value) $("settings").open = true;
  applyMode();
  // The pane stays open while other items are selected, and the same pane can show a received email or an email
  // being written (e.g. an inline reply), so the mode is re-checked every time the item changes.
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { applyMode(); if (sending) { pendingRefresh = true; return; } refreshForItem(false); });
  await refreshForItem(false);
});
