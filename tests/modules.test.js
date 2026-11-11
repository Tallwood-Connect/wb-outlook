// Checks the add-in's helper files and that the main scripts at least parse. Run: node tests/modules.test.js
const fs = require("fs"), path = require("path"), vm = require("vm");
const root = path.join(__dirname, "..");
let bad = 0;
const check = (name, ok, detail = "") => { if (!ok) { bad++; console.log("FAIL", name, detail); } else console.log("ok  ", name); };

for (const f of ["taskpane.js", "launchevent.js", "thread.js", "mime.js", "eml.js", "channel.js", "pdf.js"]) {
  try { new vm.Script(fs.readFileSync(path.join(root, f), "utf8"), { filename: f }); check(`${f} parses`, true); }
  catch (e) { check(`${f} parses`, false, e.message); }
}
// the panel version must match everywhere (the bump rule)
const v = /APP_VERSION = "([^"]+)"/.exec(fs.readFileSync(path.join(root, "taskpane.js"), "utf8"))[1];
const html = fs.readFileSync(path.join(root, "taskpane.html"), "utf8");
check(`taskpane.html data-v is ${v}`, new RegExp(`data-v="${v.replace(/\./g, "\\.")}"`).test(html));
const stale = [...html.matchAll(/\?v=([0-9.]+)/g)].map((m) => m[1]).filter((x) => x !== v);
check("every ?v= in taskpane.html equals the panel version", stale.length === 0, stale.join(","));

const g = globalThis;
new vm.Script(fs.readFileSync(path.join(root, "thread.js"), "utf8")).runInThisContext();
const T = g.WBThread;
check("thread.js exposes belongs", typeof T.belongs === "function");
const meta = { v: 1, entries: [{ id: "a", msgId: "<1@x>", convId: "C1", refs: [] }] };
check("same email saved again belongs", T.belongs(meta, { id: "a" }));
check("reply with unreadable headers joins by conversation id", T.belongs(meta, { id: "b", convId: "C1", msgId: "", refs: [] }));
check("different conversation id does not join by itself", !T.belongs(meta, { id: "c", convId: "C2", msgId: "", refs: [], quoteCount: 0, quoteRatio: 0 }));
check("reply that cites the earlier message joins", T.belongs(meta, { id: "d", msgId: "<2@x>", inReplyTo: "<1@x>", refs: ["<1@x>"] }));
check("unrelated email does not join", !T.belongs(meta, { id: "e", msgId: "<9@x>", refs: [], quoteCount: 0, quoteRatio: 0 }));

new vm.Script(fs.readFileSync(path.join(root, "mime.js"), "utf8")).runInThisContext();
const eml = Buffer.from("From: a@x\r\nMessage-ID: <m1@x>\r\nIn-Reply-To: <m0@x>\r\nReferences: <r1@x>\r\n <r2@x>\r\nSubject: Hi\r\n\r\nbody");
const h = g.WBMime.headers(new Uint8Array(eml));
check("mime.js reads Message-ID", h.msgId === "<m1@x>", JSON.stringify(h));
check("mime.js reads In-Reply-To", h.inReplyTo === "<m0@x>", JSON.stringify(h));
check("mime.js reads a folded References line", /<r1@x>\s+<r2@x>/.test(h.references || ""), JSON.stringify(h));

console.log(bad ? `${bad} failed` : "all passed");
process.exit(bad ? 1 : 0);
