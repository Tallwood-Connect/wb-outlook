/* Runs when the user clicks Send on a new email (Outlook "Smart Alerts" OnMessageSend).
   If the person has switched on "Ask me to save sent emails", it offers to save the email to Wunderbuild first.
   Safety rule: anything unexpected lets the email send. This must never trap an email. */
function onMessageSendHandler(event) {
  let finished = false;
  const done = (opts) => { if (finished) return; finished = true; try { event.completed(opts); } catch (e) { /* nothing more can be done */ } };
  try {
    const rs = Office.context.roamingSettings;
    if (!rs || rs.get("wb_ask") !== true) return done({ allowEvent: true }); // switched off (the default)
    const item = Office.context.mailbox.item;
    if (!item || typeof item.loadCustomPropertiesAsync !== "function") return done({ allowEvent: true });
    item.loadCustomPropertiesAsync((r) => {
      try {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return done({ allowEvent: true });
        if (r.value.get("wbSaved") === true) return done({ allowEvent: true }); // already saved from the panel
        done({
          allowEvent: false,
          errorMessage: "Do you want to save this email to Wunderbuild before it is sent? Choose Send Anyway to send without saving.",
          cancelLabel: "Save to Wunderbuild",
          commandId: "wbComposeButton",
        });
      } catch (e) { done({ allowEvent: true }); }
    });
    setTimeout(() => done({ allowEvent: true }), 8000); // never hold up sending if Outlook does not answer
  } catch (e) { done({ allowEvent: true }); }
}
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
