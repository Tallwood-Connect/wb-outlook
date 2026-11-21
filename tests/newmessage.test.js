// offline test for launchevent.js new-email handler (the "#wb-<code>" hand-over from Publish to Wunderbuild)
const fs = require("fs"), path = require("path");
const code = fs.readFileSync(path.join(__dirname, "..", "launchevent.js"), "utf8") + "\n;global.__h = onNewMessageComposeHandler;";
let bad = 0;
const check = (ok, msg) => { console.log(`${ok ? "ok  " : "FAIL"} ${msg}`); if (!ok) bad++; };

async function run(opts) {
  const store = Object.assign({ wb_base: "https://wb.example.test/api", wb_key: opts.key || "wbt_secretsecretsecret" }, opts.store || {});
  const seen = { reads: 0, fetches: [], subjectSet: null, notes: [] };
  global.localStorage = { getItem: (k) => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); }, removeItem: (k) => { delete store[k]; } };
  global.window = { addEventListener() {} };
  global.setTimeout = (f, ms) => { if (ms >= 100000) return 0; Promise.resolve().then(f); return 0; }; // the 170 s guard never fires; short waits run at once
  const ok = (cb, v) => cb({ status: "ok", value: v });
  global.fetch = async (url, init) => { seen.fetches.push({ url, key: init && init.headers && init.headers["x-api-key"] });
    return opts.status && opts.status !== 200 ? { ok: false, status: opts.status } : { ok: true, status: 200, json: async () => ({ subject: "Review", to: ["a@x.test"], html: "<p>hi</p>", files: [] }) }; };
  global.Office = { AsyncResultStatus: { Succeeded: "ok", Failed: "fail" }, CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { ItemNotificationMessageType: { ErrorMessage: "e", InformationalMessage: "i" } },
    actions: { associate() {} }, onReady() {}, context: { roamingSettings: { get: () => "" }, mailbox: { item: {
      subject: { getAsync: (cb) => { const v = opts.subjects[Math.min(seen.reads, opts.subjects.length - 1)]; seen.reads++; ok(cb, v); },
        setAsync: (v, cb) => { seen.subjectSet = v; ok(cb); } },
      to: { setAsync: (v, cb) => ok(cb) },
      body: { getAsync: (t, cb) => ok(cb, "sig"), prependAsync: (h, o, cb) => ok(cb) },
      isClientSignatureEnabledAsync: (cb) => ok(cb, false),
      notificationMessages: { replaceAsync: (k, d, cb) => { seen.notes.push(d.message); cb && cb(); } },
    } } } };
  new Function("global", code)(global);
  const why = await new Promise((resolve) => { const log = console.log; console.log = (m) => { const s = String(m); if (/new email: /.test(s)) { console.log = log; resolve(s); } }; global.__h({ completed: () => {} }); });
  return { seen, why, log: JSON.parse(store.wb_le_log || "[]") };
}

(async () => {
  const origLog = console.log;
  let r = await run({ subjects: ["Review #wb-ABCDEF123456"] });
  console.log = origLog;
  check(r.seen.fetches.length === 1 && r.seen.fetches[0].url === "https://wb.example.test/mail/ABCDEF123456", "tag at once: collects the email from the Worker's /mail");
  check(r.seen.reads === 1 && /filled in/.test(r.why), "tag at once: reads the subject once and fills the draft in");
  check(r.seen.subjectSet === "Review", "the draft's subject is replaced by the email's own subject");

  r = await run({ subjects: ["", "", "Review #wb-ABCDEF123456"] });
  console.log = origLog;
  check(r.seen.reads === 3 && r.seen.fetches.length === 1, "empty subject, tag arrives on the third look: collected");
  check(r.log.some((l) => /subject arrived after/.test(l.m)), "the log says the subject arrived late");

  r = await run({ subjects: [""] });
  console.log = origLog;
  check(r.seen.reads === 11 && r.seen.fetches.length === 0, "subject stays empty: looks 10 more times (about 3 s), then leaves the email alone");
  check(/not from Publish/.test(r.why) && r.log.some((l) => /still empty after/.test(l.m)), "an ordinary new email finishes and says why");

  r = await run({ subjects: ["Lunch on Friday"] });
  console.log = origLog;
  check(r.seen.reads === 1 && r.seen.fetches.length === 0, "a subject with other text and no tag is not waited for");

  r = await run({ subjects: ["", "Lunch"] });
  console.log = origLog;
  check(r.seen.reads === 2 && r.seen.fetches.length === 0, "text arrives with no tag: stops looking at once");

  r = await run({ subjects: ["Review #wb-ABCDEF123456"], status: 404 });
  console.log = origLog;
  check(/error: couldn't collect the email \(404\)/.test(r.why) && r.seen.notes.some((n) => /404/.test(n)), "a refused collect shows the status in the draft's info bar");
  check(r.log.some((l) => /collect answered 404/.test(l.m)), "the log keeps the answer status");

  const all = JSON.stringify(r.log);
  check(!/wbt_secret|secretsecret/.test(all), "the log never contains the token");
  check(/person token/.test(all), "the log says whether a person or shared token is used");
  check(r.log.length <= 40, "the log keeps at most 40 lines");
  r = await run({ subjects: ["Lunch on Friday"] });
  console.log = origLog;
  check(!/Lunch|Friday/.test(JSON.stringify(r.log)), "the log holds no subject text");

  if (bad) { console.log(`${bad} failed`); process.exit(1); }
  console.log("all passed");
})();
