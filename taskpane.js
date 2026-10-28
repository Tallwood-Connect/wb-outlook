/* Send to Wunderbuild: Outlook task pane. Uses the Wunderbuild public API (beta). */
const DEFAULT_BASE = "https://api.wunderbuild.com/api";
const $ = (id) => document.getElementById(id);
const store = {
  get: (k) => { try { return localStorage.getItem(k); } catch { return null; } },
  set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } },
};
let job = null; // { _id, name }

function setStatus(msg, kind = "") { const s = $("status"); s.textContent = msg; s.className = kind; }
function base() { return ($("base").value.trim() || DEFAULT_BASE).replace(/\/+$/, ""); }
function refreshSend() { $("send").disabled = !(job && $("name").value.trim()); }

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

/* ---------- job search ---------- */
let timer;
$("q").addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(searchJobs, 300); });

async function searchJobs() {
  const q = $("q").value.trim();
  const list = $("jobs");
  list.textContent = "";
  if (!q) return;
  try {
    setStatus("Searching…");
    const data = await api(`/jobs?page=1&limit=15&search=${encodeURIComponent(q)}`);
    setStatus(data.result.length ? "" : "No jobs found.");
    for (const j of data.result) {
      const li = document.createElement("li");
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = `${j.jobNumber ? "#" + j.jobNumber + " " : ""}${j.name || "Untitled job"}`;
      if (j.address) { const s = document.createElement("small"); s.textContent = j.address; b.appendChild(s); }
      b.addEventListener("click", () => chooseJob(j));
      li.appendChild(b); list.appendChild(li);
    }
  } catch (e) { setStatus(e.message, "err"); }
}

async function chooseJob(j) {
  job = j;
  store.set("wb_job", JSON.stringify({ _id: j._id, name: j.name, jobNumber: j.jobNumber }));
  $("jobs").textContent = "";
  $("q").value = "";
  $("chosen").textContent = `Job: ${j.jobNumber ? "#" + j.jobNumber + " " : ""}${j.name}`;
  refreshSend();
  await loadFolders();
}

/* ---------- folders ---------- */
async function loadFolders() {
  const sel = $("folder");
  sel.disabled = true;
  sel.innerHTML = "";
  sel.add(new Option("Job documents (top level)", ""));
  try {
    const root = await api(`/documents/job/${job._id}/folders`);
    const walk = (folders, prefix) => {
      for (const f of folders || []) {
        sel.add(new Option(prefix + f.name, f._id));
        walk(f.subFolders, prefix + f.name + " / ");
      }
    };
    walk(root.subFolders, "");
    // Default to the Correspondence folder (also matches it when nested, e.g. "Admin / Correspondence").
    const corr = [...sel.options].find((o) => /(^|\/ )correspondence$/i.test(o.text.trim()));
    if (corr) sel.value = corr.value;
  } catch (e) { setStatus(e.message, "err"); }
  sel.disabled = false;
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
  const item = Office.context.mailbox.item;
  if (!item) return "";
  const d = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
  const day = d.toISOString().slice(0, 10);
  const subject = (item.subject || "No subject").replace(/[\\/:*?"<>|\r\n]+/g, " ").trim().slice(0, 100);
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
function emailMeta() {
  const item = Office.context.mailbox.item;
  const atts = (item.attachments || []).filter((a) => !a.isInline).map((a) => `${a.name} (${Math.max(1, Math.round(a.size / 1024))} KB)`);
  return [
    ["From:", fmtAddr(item.from)],
    ["To:", fmtList(item.to)],
    ["Cc:", fmtList(item.cc)],
    ["Date:", item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString("en-AU") : ""],
    ["Attach:", atts.join("; ")],
  ];
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
async function ensureOutlookFolder(parentId) {
  const detail = await api(parentId ? `/documents/job/${job._id}/folders/${parentId}` : `/documents/job/${job._id}/folders`);
  const hit = (detail.subFolders || []).find((f) => (f.name || "").trim().toLowerCase() === "outlook files");
  if (hit) return hit._id;
  const created = await api(`/documents/job/${job._id}/folders`, {
    method: "POST",
    body: JSON.stringify({ name: "Outlook Files", ...(parentId ? { parentId } : {}) }),
  });
  return created._id;
}

$("send").addEventListener("click", async () => {
  const btn = $("send");
  btn.disabled = true;
  try {
    const baseName = $("name").value.trim().replace(/\.(eml|pdf|msg)$/i, "");
    const folderId = $("folder").value;
    const folderLabel = $("folder").selectedOptions[0].text;

    setStatus("Reading email…");
    const item = Office.context.mailbox.item;
    const [bodyText, emlB64] = await Promise.all([getBodyText(), getEml()]);
    const pdf = emailToPdf({ subject: item.subject, meta: emailMeta(), body: bodyText });
    const eml = b64ToBytes(emlB64);

    setStatus("Saving PDF…");
    await uploadFile(`${baseName}.pdf`, "application/pdf", pdf, folderId);

    try {
      setStatus("Saving original email to Outlook Files…");
      const outlookFolderId = await ensureOutlookFolder(folderId);
      await uploadFile(`${baseName}.eml`, "message/rfc822", eml, outlookFolderId);
    } catch (e) {
      setStatus(`PDF saved to ${folderLabel}, but the original email was not saved: ${e.message}`, "err");
      refreshSend();
      return;
    }
    setStatus(`Saved. PDF in ${folderLabel}; original email in ${folderLabel} / Outlook Files.`, "ok");
  } catch (e) {
    setStatus(e.message, "err");
  }
  refreshSend();
});

/* ---------- init ---------- */
$("name").addEventListener("input", refreshSend);
["key", "base"].forEach((id) => $(id).addEventListener("change", () => store.set("wb_" + id, $(id).value.trim())));

Office.onReady(async () => {
  $("key").value = store.get("wb_key") || "";
  $("base").value = store.get("wb_base") || DEFAULT_BASE;
  if (!$("key").value) $("settings").open = true;
  $("name").value = suggestName();
  // Keep the pane pinned: refresh the file name when another email is selected.
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { $("name").value = suggestName(); setStatus(""); refreshSend(); });
  try {
    const saved = JSON.parse(store.get("wb_job") || "null");
    if (saved && saved._id) await chooseJob(saved);
  } catch { /* ignore */ }
  refreshSend();
});
