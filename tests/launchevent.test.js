// offline test for launchevent.js send-time prompt filter
const fs = require("fs");
function run(opts) {
  const store = Object.assign({}, opts.store || {});
  const out = { completed: null };
  global.localStorage = { getItem: (k) => (k in store ? store[k] : null), setItem() {}, removeItem() {} };
  global.window = { addEventListener() {} };
  global.setTimeout = (f, ms) => 0;
  global.Office = { AsyncResultStatus: { Succeeded: "ok", Failed: "fail" }, context: { roamingSettings: { get: (k) => (k === "wb_ask" ? opts.ask !== false : "") },
    mailbox: { item: { conversationId: opts.conv, loadCustomPropertiesAsync: (cb) => cb({ status: "ok", value: { get: (k) => (k === "wbSaved" ? !!opts.saved : undefined) } }),
      subject: opts.noSubject ? undefined : { getAsync: (cb) => cb(opts.subjFail ? { status: "fail" } : { status: "ok", value: opts.subject }) } } } } };
  const code = fs.readFileSync(""+require("path").join(__dirname,"..","launchevent.js")+"", "utf8") + "\n;global.__h = onMessageSendHandler;";
  new Function("global", code)(global);
  global.__h({ completed: (o) => { out.completed = o; } });
  return out.completed;
}
const jobs = { at: 1, jobs: [ { jobNumber: 1026, name: "Farrell - D" }, { jobNumber: 987, name: "Hann" }, { jobNumber: "01053", name: "Smith Residence" }, { jobNumber: 1100, name: "Jo" } ] };
const S = { wb_jobs_v2: JSON.stringify(jobs) };
const t = [
  ["J- ref", { subject: "J-01026 - Quote", store: S }, true],
  ["j- lower", { subject: "re: j-1026", store: S }, true],
  ["bracket J-", { subject: "[J-01067 - Smith] hello", store: S }, true],
  ["bare 4 digit no longer asks", { subject: "Query on 1026", store: S }, false],
  ["job word + 5 digit padded", { subject: "Job 01026 shaft", store: S }, true],
  ["#number asks", { subject: "plan #1053 rev", store: S }, true],
  ["bare 1053 no", { subject: "plan 1053 rev", store: S }, false],
  ["No. number asks", { subject: "Quote No. 1026", store: S }, true],
  ["4 digit no match", { subject: "Invoice 2026", store: S }, false],
  ["3 digit never", { subject: "unit 987", store: S }, false],
  ["number inside longer", { subject: "ref 110265", store: S }, false],
  ["one-word name alone no", { subject: "Farrell shaft pit", store: S }, false],
  ["one-word name + job word", { subject: "Farrell job shaft pit", store: S }, true],
  ["one-word name + own number", { subject: "Farrell 1026 shaft", store: S }, true],
  ["name word boundary", { subject: "Farrelly lunch", store: S }, false],
  ["name 2 words", { subject: "Re: Smith Residence drawings", store: S }, true],
  ["short name ignored", { subject: "Jo says hi", store: S }, false],
  ["no signal", { subject: "Lunch on Friday", store: S }, false],
  ["thread saved", { subject: "Lunch", conv: "C1", store: Object.assign({ wb_conv_C1: JSON.stringify({ at: Date.now() }) }, S) }, true],
  ["reply record", { subject: "Lunch", conv: "C2", store: Object.assign({ wb_conv_reply_C2: JSON.stringify({ at: Date.now() }) }, S) }, true],
  ["old thread expired", { subject: "Lunch", conv: "C3", store: Object.assign({ wb_conv_C3: JSON.stringify({ at: Date.now() - 100 * 864e5 }) }, S) }, false],
  ["already saved sends", { subject: "J-01026", saved: true, store: S }, false],
  ["setting off sends", { subject: "J-01026", ask: false, store: S }, false],
  ["subject unreadable asks", { subjFail: true, store: S }, true],
  ["no subject api asks", { noSubject: true, store: S }, true],
  ["no job list, no signal asks", { subject: "Lunch", store: {} }, true],
  ["bad job list asks", { subject: "Lunch", store: { wb_jobs_v2: "{bad" } }, true],
];
let bad = 0;
for (const [n, o, ask] of t) { const r = run(o); const got = r && r.allowEvent === false; if (got !== ask) { bad++; console.log("FAIL", n, JSON.stringify(r)); } else console.log("ok  ", n); }
console.log(bad ? bad + " failed" : "all passed"); if (bad) process.exit(1);
