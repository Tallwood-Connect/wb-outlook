/* Runs when the user clicks Send on a new email (Outlook "Smart Alerts" OnMessageSend).
   If the person has switched on "Ask me to save sent emails", it offers to save the email to Wunderbuild first.
   Safety rule: anything unexpected lets the email send. This must never trap an email.
   Messages starting with [WB] go to the browser console, to help find problems. */
const WB_LE_VERSION = "1.24";
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
        done({
          allowEvent: false,
          errorMessage: "Would you like to save this email to Wunderbuild?",
          cancelLabel: "Send to Wunderbuild",
          commandId: "wbComposeButton",
        }, "not saved yet");
      } catch (e) { done({ allowEvent: true }, `error: ${e && e.message}`); }
    });
  } catch (e) { done({ allowEvent: true }, `error: ${e && e.message}`); }
}

try {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler); // must happen as soon as the script loads
  wbLog("handler registered");
} catch (e) { wbLog(`could not register handler: ${e && e.message}`); }
try { Office.onReady(() => wbLog("Office is ready")); } catch (e) { /* logging only; not needed for sending */ }
