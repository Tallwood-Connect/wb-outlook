/* Two automatic handlers (no button press needed):
   1. OnMessageSend (Outlook "Smart Alerts"): if the person has switched on "Ask me to save sent emails", it
      offers to save the email to Wunderbuild first. Safety rule: anything unexpected lets the email send.
      This must never trap an email.
   2. OnNewMessageCompose: when Archicad's "Publish to Wunderbuild" opens a new email whose subject ends in
      "#wb-<code>", it fills that email in (recipients, subject, message above the Outlook signature, the
      published files attached from Wunderbuild). Every other new email is left untouched.
   Messages starting with [WB] go to the browser console, to help find problems. */
const WB_LE_VERSION = "1.27";
const wbLog = (m) => { try { console.log(`[WB ${WB_LE_VERSION}] ${m}`); } catch (e) { /* no console */ } };
wbLog("send-time script loaded");
try { window.addEventListener("error", (e) => wbLog(`script error: ${e && e.message}`)); } catch (e) { /* not a browser page */ }

function onMessageSendHandler(event) {
  wbLog("handler started");
  let finished = false;
  const done = (opts, why) => {
    if (finished) return;
    finished = true;
    wbLog(`completing: ${opts.allowEvent ? "allow send" : "ask to save"} (${why})`);
    try { event.completed(opts); } catch (e) { wbLog(`event.completed failed: ${e && e.message}`); }
  };
  try {
    const rs = Office.context.roamingSettings;
    if (!rs || rs.get("wb_ask") !== true) return done({ allowEvent: true }, "setting is off");
    const item = Office.context.mailbox.item;
    if (!item || typeof item.loadCustomPropertiesAsync !== "function") return done({ allowEvent: true }, "no draft properties");
    setTimeout(() => done({ allowEvent: true }, "timed out waiting for Outlook"), 8000); // never hold up sending
    item.loadCustomPropertiesAsync((r) => {
      try {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return done({ allowEvent: true }, "could not read draft properties");
        if (r.value.get("wbSaved") === true) return done({ allowEvent: true }, "already saved");
        const block = () => done({
          allowEvent: false,
          errorMessage: "Would you like to save this email to Wunderbuild?",
          cancelLabel: "Send to Wunderbuild",
          commandId: "wbComposeButton",
        }, "not saved yet");
        block();
      } catch (e) { done({ allowEvent: true }, `error: ${e && e.message}`); }
    });
  } catch (e) { done({ allowEvent: true }, `error: ${e && e.message}`); }
}

/* ---------- 2. Fill in an email prepared by Publish to Wunderbuild ---------- */
const WB_MAIL_TAG = /\s*#wb-([A-Za-z0-9]{6,16})\s*$/;
const WB_MAX_ATTACH = 20 * 1024 * 1024; // same limit as the task pane
// The API settings saved by the task pane (this page shares its storage; roaming settings are the fallback).
function wbSetting(k) {
  try { const v = localStorage.getItem(k); if (v) return v; } catch (e) { /* no storage here */ }
  try { return Office.context.roamingSettings.get(k) || ""; } catch (e) { return ""; }
}
function wbCall(fn) {
  return new Promise((resolve, reject) => fn((r) => (r.status === Office.AsyncResultStatus.Succeeded
    ? resolve(r.value) : reject(new Error((r.error && r.error.message) || "Outlook refused")))));
}
function wbB64(bytes) {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
function wbNote(item, text, isError) {
  try {
    item.notificationMessages.replaceAsync("wbPublish", {
      type: isError ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: String(text).slice(0, 150), icon: "Icon.16x16", persistent: !isError,
    }, () => {});
  } catch (e) { /* notifications unavailable */ }
}
const wbSleep = (ms) => new Promise((r) => setTimeout(r, ms));
// Outlook adds the user's signature to a new email itself. Wait for it (up to 4 s) so the message can go above it.
async function wbWaitForSignature(item) {
  let auto = true;
  try { auto = await wbCall((cb) => item.isClientSignatureEnabledAsync(cb)); } catch (e) { /* older Outlook: assume yes */ }
  if (!auto) return;
  for (let i = 0; i < 16; i++) {
    try {
      const text = await wbCall((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
      if (String(text || "").trim()) return;
    } catch (e) { return; }
    await wbSleep(250);
  }
}
async function wbFetchBytes(root, key, url) {
  let res;
  try { res = await fetch(url); } catch (e) { res = await fetch(root + "/download", { headers: { "x-api-key": key, "x-download-url": url } }); }
  if (!res.ok) throw new Error(`download failed (${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

function onNewMessageComposeHandler(event) {
  let finished = false;
  const done = (why) => {
    if (finished) return;
    finished = true;
    wbLog(`new email: ${why}`);
    try { event.completed(); } catch (e) { wbLog(`event.completed failed: ${e && e.message}`); }
  };
  setTimeout(() => done("stopped waiting"), 170000);
  const item = Office.context.mailbox.item;
  (async () => {
    const subject = await wbCall((cb) => item.subject.getAsync(cb));
    const tag = WB_MAIL_TAG.exec(subject || "");
    if (!tag) return done("not from Publish to Wunderbuild");          // every other new email: untouched
    const base = (wbSetting("wb_base") || "").trim().replace(/\/+$/, "");
    const key = (wbSetting("wb_key") || "").trim();
    if (!base || !key) {
      wbNote(item, "Publish to Wunderbuild: open the Wunderbuild add-in once so it can fill in these emails.", true);
      return done("no API settings");
    }
    const root = base.replace(/\/api$/, "");
    const res = await fetch(`${root}/mail/${tag[1]}`, { headers: { "x-api-key": key } });
    if (!res.ok) throw new Error(`couldn't collect the email (${res.status})`);
    const mail = await res.json();
    wbNote(item, "Publish to Wunderbuild is filling in this email…", false);
    await wbCall((cb) => item.subject.setAsync(mail.subject || subject.replace(WB_MAIL_TAG, ""), cb));
    if (Array.isArray(mail.to) && mail.to.length) await wbCall((cb) => item.to.setAsync(mail.to, cb));
    await wbWaitForSignature(item);
    await wbCall((cb) => item.body.prependAsync(mail.html || "", { coercionType: Office.CoercionType.Html }, cb));
    const failed = [];
    for (const f of Array.isArray(mail.files) ? mail.files : []) {
      try {
        const info = await (await fetch(`${base}/documents/job/${f.job}/${f.doc}/download-url`,
          { headers: { "x-api-key": key, "Content-Type": "application/json" } })).json();
        const bytes = await wbFetchBytes(root, key, info.url);
        if (bytes.length > WB_MAX_ATTACH) throw new Error("too large to attach");
        await wbCall((cb) => item.addFileAttachmentFromBase64Async(wbB64(bytes), f.name || info.name, { isInline: false }, cb));
      } catch (e) { failed.push(`${f.name}: ${e.message}`); }
    }
    wbNote(item, failed.length ? `Not attached: ${failed.join("; ")}` :
      "Prepared by Publish to Wunderbuild. Check it, then Send.", failed.length > 0);
    done(failed.length ? "filled in, some files not attached" : "filled in");
  })().catch((e) => {
    wbNote(item, `Publish to Wunderbuild couldn't fill in this email: ${e && e.message}`, true);
    done(`error: ${e && e.message}`);
  });
}

try {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  wbLog("new-email handler registered");
} catch (e) { wbLog(`could not register new-email handler: ${e && e.message}`); }
try {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler); // must happen as soon as the script loads
  wbLog("handler registered");
} catch (e) { wbLog(`could not register handler: ${e && e.message}`); }
try { Office.onReady(() => wbLog("Office is ready")); } catch (e) { /* logging only; not needed for sending */ }
